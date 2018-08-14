const { ApolloServer, gql } = require ('apollo-server');

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
    type Query {
        books: [Book]
    }
`;

const resolvers = {
    Query: {
        books: () => books,
    }
}

const server = new ApolloServer({typeDefs,resolvers});

server.listen().then(({url})=>{
    console.log(`Server ready at ${url}`);
});