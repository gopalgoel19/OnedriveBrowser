const { ApolloServer, gql } = require ('apollo-server');
const { GraphQLServer } = require('graphql-yoga')
const fetch = require('node-fetch')

const baseURL = `https://graph.microsoft.com/v1.0/me`

const books = [
    {
        title: "Godfather"
    },
    {
        title: "White Tiger"
    }
];

const typeDefs = gql`
    type Book {
        title: String
    }

    type item {
        id: String
    }

    type Query {
        books: [Book]
        items: [item]
    }
`;


const resolvers = {
    Query: {
        items: async (parent,args,context) => {
            // console.log(context.authorization);
            return await fetch(`${baseURL}/drive/root/children`, {
                method: "GET",
                headers: {authorization: context.authorization}
            })
            .then(res => res.json())
            .then(json => {
                let items = [];
                json['value'].forEach((obj)=>{
                    items.push({'id': obj.id});
                })
                console.log(items);
                return items;
            });
        },
        books: () => (books)
    }
}

const server = new ApolloServer({
    typeDefs,
    resolvers,
    context: ({req})=> ({
        authorization: req.headers.authorization
    })
});

server.listen().then(({url})=>{
    console.log(`Server ready at ${url}`);
});